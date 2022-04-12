<?php
namespace Newageerp\SfImport\Controller;

use Doctrine\ORM\EntityManagerInterface;
use Doctrine\Persistence\ObjectRepository;
use Exception;
use Newageerp\SfAuth\Service\AuthService;
use Newageerp\SfUservice\Controller\UControllerBase;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Symfony\Component\EventDispatcher\EventDispatcherInterface;
use Symfony\Component\HttpFoundation\File\UploadedFile;
use Symfony\Component\HttpFoundation\JsonResponse;
use Symfony\Component\HttpFoundation\Request;
use Symfony\Component\Routing\Annotation\Route;
use OpenApi\Annotations as OA;

/**
 * @Route(path="/app/nae-core/import")
 */
class ImportController extends UControllerBase
{
    protected array $letters = [];

    public function __construct(EntityManagerInterface $em, EventDispatcherInterface $eventDispatcher)
    {
        parent::__construct($em, $eventDispatcher);

        $this->letters = ['', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'X', 'Y', 'Z'];
    }

    /**
     * @Route ("/mainImport", methods={"POST"})
     * @param Request $request
     * @param EntityManagerInterface $entityManager
     * @return JsonResponse
     * @OA\Post (operationId="NAEUImport")
     * @throws Exception
     */
    public function mainImport(Request $request, EntityManagerInterface $entityManager): JsonResponse
    {
        $request = $this->transformJsonBody($request);

        $user = $this->findUser($request);
        if (!$user) {
            throw new Exception('Invalid user');
        }
        AuthService::getInstance()->setUser($user);

        $schema = $request->get('schema');
        $properties = $this->getPropertiesForSchema($schema);

        $className = $this->convertSchemaToEntity($schema);

        /**
         * @var ObjectRepository $repository
         */
        $repository = $entityManager->getRepository($className);

        $files = $request->files;

        $keysForSave = [];
        $types = [];
        /**
         * @var UploadedFile $file
         */
        foreach ($files as $file) {
            $spreadsheet = IOFactory::load($file->getPathname());

            $sheet = $spreadsheet->getActiveSheet();

            $cols = $sheet->getCell('G1')->getValue();

            $highestRow = $sheet->getHighestRow();

            for ($i = 1; $i <= $cols; $i++) {
                $keyCol = $this->letters[$i] . '2';
                $keysForSave[] = $sheet->getCell($keyCol)->getValue();
            }

            for ($i = 4; $i <= $highestRow; $i++) {
                $idCol = 'A' . $i;

                $id = $sheet->getCell($idCol)->getValue();
                $element = $repository->find($id);

                if ($element) {
                    foreach ($keysForSave as $key => $fieldKey) {
                        if ($fieldKey) {
                            $type = null;
                            $format = null;
                            if (isset($properties[$fieldKey], $properties[$fieldKey]['type']) && $properties[$fieldKey]['type']) {
                                $type = $properties[$fieldKey]['type'];
                            }
                            if (isset($properties[$fieldKey], $properties[$fieldKey]['format']) && $properties[$fieldKey]['format']) {
                                $format = $properties[$fieldKey]['format'];
                            }

                            $types[$fieldKey] = [
                                'type' => $type,
                                'format' => $format
                            ];

                            $keyCol = $this->letters[($key + 1)] . '' . $i;
                            $val = $sheet->getCell($keyCol)->getCalculatedValue();
                            if ($type === 'number' && $format === 'float') {
                                $val = (float)$val;
                            }

                            $setter = 'set' . $fieldKey;
                            $element->$setter($val);
                        }
                    }
                }
            }
            $entityManager->flush();
        }


        return $this->json(['success' => 1, 'types' => $types]);
    }
}